import discord
from discord.ext import commands
from core.classes import Cog_Extension

class Main(Cog_Extension):
    


    
    @commands.command()
    async def invite(self,ctx,max_age=0,unit="s",max_uses=0,reason='bot invite',):
        if unit=="d" or unit=="day" or unit=="days": max_age=86400*max_age
        if unit=="h" or unit=="hour" or unit=="hours": max_age=3600*max_age
        if unit=="m" or unit=="min" or unit=="minute" or unit=="minutes": max_age=60*max_age
        invitelink = await ctx.channel.create_invite(max_age=max_age,max_uses=max_uses,unique=False,reason=reason)
        await ctx.channel.send(invitelink) 

    @commands.command()
    async def post(self,ctx,text):
        await ctx.channel.purge(limit=1)
        await ctx.channel.send(text)


    #Error Handler
    @commands.Cog.listener()
    async def on_commands_error(self,ctx,error):
        if hasattr(ctx.commands,'on_error'):
            return
        if isinstance(error,commands.errors.CommandNotFound):
            await ctx.send(f'指令[ {ctx.commamds} ]不存在,使用$help確認指令!')
        if isinstance(error,commands.errors.ConversionError):
            await ctx.send('資料傳換錯誤')
        if isinstance(error,commands.errors.MissingRequiredArgument):
            await ctx.send('所需指令參數遺失')
        if isinstance(error,commands.errors.ArgumentParsingError):
            await ctx.send('無法判斷輸入指令')
        if isinstance(error,commands.errors.DisabledCommand):
            await ctx.send('指令已發生禁用')
        if isinstance(error,commands.errors.MessageNotFound): 
            await ctx.send('訊息已遭到刪除')
        else: await ctx.send('無法識別的錯誤')


def setup(bot):
    bot.add_cog(Main(bot))